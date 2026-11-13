import { useEffect, useState } from 'react';
import type { Editor } from '@docx-editor.dev/core/contracts/editor';
import type { ReviewItemPlacement } from '@docx-editor.dev/core/contracts/editor';

export interface ReviewState {
  readonly items: readonly ReviewItemPlacement[];
  readonly activeKey: string | null;
}

const EMPTY: ReviewState = { items: [], activeKey: null };

/**
 * The editor's review queue and the activated item. Re-derives only when the engine's
 * review counter moves or the selection changes the activated card.
 */
export function useReviewItems(editor: Editor | null): ReviewState {
  const [state, setState] = useState<ReviewState>(EMPTY);
  useEffect(() => {
    if (!editor) {
      setState(EMPTY);
      return undefined;
    }
    let revision = -1;
    let items: readonly ReviewItemPlacement[] = [];
    let activeKey: string | null = null;
    const read = (force = false) => {
      const nextRevision = editor.getReviewRevision();
      const nextActive = editor.getActivatedReviewItemKey();
      if (!force && nextRevision === revision && nextActive === activeKey) return;
      if (force || nextRevision !== revision) {
        revision = nextRevision;
        items = editor.getReviewItems({ placement: false });
      }
      activeKey = nextActive;
      setState({ items, activeKey });
    };
    read();
    // Some comment commands emit change before the review revision counter advances.\n    // Force a queue refresh on document changes so new comments appear immediately.\n    const offChange = editor.on('change', () => read(true));
    const offSelection = editor.on('selectionChange', read);
    const offResolved = editor.on('reviewChangesResolved', () => read(true));
    return () => {
      offChange();
      offSelection();
      offResolved();
    };
  }, [editor]);
  return state;
}
