/**
 * פענוח `otzaria://` — הצד ההופכי ל-at-mention.ts:buildRefHref.
 *
 * מכוסות שלוש הצורות שהתוסף עצמו כותב; כל השאר מוחזר `null` ונשאר באחריות
 * אוצריא (הראוטר שלה מכיר עוד עשרות יעדים, ואין טעם לשכפל אותם כאן).
 */

export type OtzariaLinkTarget =
  | { kind: 'book'; id: number; index: number }
  | { kind: 'pdf'; id: number; index: number }
  | { kind: 'detection'; query: string };

/** `null` = לא קישור אוצריא, או צורה שאיננו מכירים. */
export function parseOtzariaLink(href: string): OtzariaLinkTarget | null {
  if (typeof href !== 'string' || !href.trim()) return null;

  let url: URL;
  try {
    url = new URL(href.trim());
  } catch {
    return null;
  }
  if (url.protocol !== 'otzaria:') return null;
  // `otzaria://open/book/42` — ה-host הוא `open`, והנתיב הוא השאר.
  if (url.hostname.toLowerCase() !== 'open') return null;

  const segments = url.pathname.split('/').filter(Boolean);
  const [kind, rawId] = segments;

  if (kind === 'detection') {
    const query = url.searchParams.get('q')?.trim() ?? '';
    return query ? { kind: 'detection', query } : null;
  }

  if (kind !== 'book' && kind !== 'pdf') return null;
  const id = Number(rawId);
  if (!Number.isInteger(id) || id <= 0) return null;

  const rawIndex = Number(url.searchParams.get('index'));
  // ברירות המחדל הן אלו של הראוטר: שורה 0 בספר טקסט, עמוד 1 ב-PDF.
  const fallback = kind === 'pdf' ? 1 : 0;
  const index = Number.isInteger(rawIndex) && rawIndex >= fallback ? rawIndex : fallback;

  return { kind, id, index };
}
