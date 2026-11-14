import type { Editor } from '@docx-editor.dev/core/contracts/editor';

/**
 * Calls `read` whenever the review queue may have changed: after the engine's change,
 * selection and resolve events (on a microtask, never during React's render), and on a light
 * poll. The poll catches a change that lands after the event that announced it — the engine
 * computes its review counter when read, and on a busy machine a programmatic edit can settle
 * after the last event. `read` must be cheap when nothing changed (compare
 * `getReviewRevision()` first). Returns the unsubscribe.
 */
export function watchReview(editor: Editor, read: () => void, pollMs = 300): () => void {
  let live = true;
  const later = () => queueMicrotask(() => live && read());
  const offs = [
    editor.on('change', later),
    editor.on('selectionChange', later),
    editor.on('reviewChangesResolved', later),
  ];
  const timer = setInterval(() => live && read(), pollMs);
  return () => {
    live = false;
    clearInterval(timer);
    for (const off of offs) off();
  };
}
