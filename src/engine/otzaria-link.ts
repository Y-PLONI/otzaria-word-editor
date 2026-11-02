/**
 * פענוח `otzaria://` — הצד ההופכי ל-at-mention.ts:buildRefHref.
 *
 * מכוסות שלוש הצורות שהתוסף עצמו כותב; כל השאר מוחזר `null` ונשאר באחריות
 * אוצריא (הראוטר שלה מכיר עוד עשרות יעדים, ואין טעם לשכפל אותם כאן).
 */

/**
 * מקור הספר בערכי חוזה התוספים (`PluginBookIdentity.sourceOf` באוצריא) — לא
 * בערכי הקישור. הקישור כותב `source=user` לספר אישי ו-`source=db:<slug>`
 * למסד מצורף; בלי הפרמטר הספר רשמי.
 */
export type OtzariaBookSource = 'library' | 'user' | 'attached';

export type OtzariaLinkTarget =
  | {
      kind: 'book';
      id: number;
      index: number;
      uid?: string;
      source?: OtzariaBookSource;
      /** `m=<טקסט>` — הדגשת הטקסט בתוך המקטע שב-`index`. */
      markText?: string;
      /** `mark` — הדגשת המקטע שב-`index` כולו. */
      markSection?: boolean;
    }
  | { kind: 'pdf'; id: number; index: number; uid?: string; source?: OtzariaBookSource }
  | { kind: 'detection'; query: string };

/**
 * `source` של הקישור, כפי שהראוטר של אוצריא קורא אותו
 * (`parseBookLinkSourceParam` ב-book_link_builder.dart).
 *
 * `undefined` כשהפרמטר חסר: קישור שהתוסף עצמו כותב (at-mention.ts:buildRefHref)
 * אינו נושא `source` גם לספר ממסד מצורף, ושליחת `library` בשמו הייתה פוסלת
 * אותו אצל המאחז כש-`uid` אינו נפתר. ערך שהראוטר אינו מכיר — גם `undefined`,
 * ואז ההתנהגות היא זו שלפני השדה.
 */
function linkSource(
  raw: string | null,
  id: number,
): { source: OtzariaBookSource; uid?: string } | undefined {
  const trimmed = raw?.trim() ?? '';
  const value = trimmed.toLowerCase();
  if (!value) return undefined;
  if (value === 'official') return { source: 'library' };
  if (value === 'user') return { source: 'user' };
  if (value.startsWith('db:')) {
    // BookSource.isValidSlug: אין לשנות את רישיות ה-slug. מסדים שונים
    // חולקים id, ולכן `attached` לבדו אינו זהות מספקת (BookIdentity.keyOf).
    const slug = trimmed.slice(3);
    if (/^[\p{L}\p{N}._-]{1,64}$/u.test(slug)) {
      return { source: 'attached', uid: `db:${slug}:${id}` };
    }
  }
  return undefined;
}

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

  // `uid` הוא המזהה היציב של הספר (ראו at-mention.ts:buildRefHref). קישור
  // ישן נכתב בלעדיו, ולכן הוא אופציונלי ולא תנאי לפענוח.
  const identity = linkSource(url.searchParams.get('source'), id);
  const uid = url.searchParams.get('uid')?.trim() || identity?.uid;
  const base = {
    id,
    index,
    ...(uid ? { uid } : {}),
    ...(identity ? { source: identity.source } : {}),
  };
  if (kind === 'pdf') return { kind, ...base };

  // `m` ו-`mark` כפי שהראוטר קורא אותם (external_uri_router.dart): `mark` הוא
  // דגל — קיים גם בלי ערך — ו-`m` ריק או רווחים בלבד אינו הדגשה. בקישור PDF
  // הראוטר אינו קורא אותם, ולכן גם כאן לא.
  const markText = url.searchParams.get('m')?.trim();
  const markSection = url.searchParams.has('mark');
  return {
    kind,
    ...base,
    ...(markText ? { markText } : {}),
    ...(markSection ? { markSection } : {}),
  };
}
