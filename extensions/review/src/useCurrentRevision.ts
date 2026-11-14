import { useEffect, useState } from 'react';
import type { Editor } from '@docx-editor.dev/core/contracts/editor';
import { watchReview } from './watch-review';

/**
 * The key of the tracked change the caret is in, as Word's Accept and Reject act on: the
 * change Next/Previous Change activated, else the one the caret sits in. `null` outside any
 * change.
 */
export function useCurrentRevision(editor: Editor | null): string | null {
  const [key, setKey] = useState<string | null>(null);
  useEffect(() => {
    if (!editor) {
      setKey(null);
      return undefined;
    }
    let seen = '';
    const read = () => {
      // Recompute only when the queue or the caret moved (the poll would otherwise re-derive
      // the whole queue every tick). The counter covers the queue and the activated item; the
      // caret can move within a line without moving it, so the selection is part of the key.
      const state = `${editor.getReviewRevision()}|${JSON.stringify(editor.snapshot().selection)}`;
      if (state === seen) return;
      seen = state;
      const items = editor.getReviewItems({ placement: false });
      const activated = editor.getActivatedReviewItemKey();
      const current =
        items.find((item) => item.kind === 'revision' && item.key === activated) ??
        items.find((item) => item.kind === 'revision' && item.isActive);
      setKey(current?.key ?? null);
    };
    read();
    return watchReview(editor, read);
  }, [editor]);
  return key;
}
