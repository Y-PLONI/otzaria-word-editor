/**
 * השלמה ממילונים סטטיים (ביטויים תלמודיים, שמות מחברים) — מוטבעים בקוד,
 * בלי RPC ובלי לטעון תוכן מהספרייה. משתמשת ישירות במנוע ההתאמה של "השלמה
 * מהספר" (`book-completion.ts`): אותו רעיון בדיוק — הקשר + מילה חלקית →
 * השלמה — רק שהמקור הוא רשימה קבועה ולא טקסט חי מהקורא. כל רשימה נבנית
 * כ"מסמך וירטואלי" (שורה לכל ערך) ומקבלת `SectionWordCache` משלה.
 *
 * מופעלת **רק** אחרי שההשלמה מהספר הפתוח בקורא לא מצאה התאמה לאותה הקשה —
 * כך סדר העדיפויות הקיים (ר' docs/smart-source-completion-plan.md) לא זז.
 *
 * ## נכס נפרד, לא `import`
 *
 * נמדד תוך כדי `check:dist`: `import` ישיר של `talmudic-phrases.json` הכניס
 * את הטקסט לתוך `assets/app.js` (ה-build הוא IIFE יחיד עם
 * `inlineDynamicImports: true`, בדיוק כמו שמתועד ב-`spellcheck-dictionary.ts`)
 * — ואחד הביטויים ("סהדי שקרי אאוגרייהו זילי") מכיל בטעות את אותה תת-מחרוזת
 * ששער `check:dist` משתמש בה כסמן לדליפת המילון התורני, ותפס את זה כשגיאת
 * false-positive. מעבר לנכס עצל, בדיוק כמו המילון וראשי-התיבות, פותר את שתי
 * הבעיות יחד: גם `app.js` לא נושא רשימות מילים חופשיות, וגם אין קונפליקט סמן.
 */
import { buildSectionCache, matchAtCursor, type BookCompletionMatch, type SectionWordCache, type TypedContext } from './book-completion';
import { STATIC_COMPLETION_FILE, STATIC_COMPLETION_GLOBAL } from './static-completion-constants';

export interface StaticSource {
  readonly name: string;
  readonly cache: SectionWordCache;
}

interface PackedStaticCompletion {
  phrases: string[];
  authors: string[];
}

function buildSource(name: string, entries: readonly string[]): StaticSource {
  return { name, cache: buildSectionCache(entries.join('\n')) };
}

function buildSources(packed: PackedStaticCompletion): StaticSource[] {
  return [
    buildSource('talmudic-phrases', packed.phrases),
    buildSource('authors', packed.authors),
  ];
}

/* ===========================================================================
 *  טעינה עצלה — אותה תבנית בדיוק כמו spellcheck-dictionary.ts/acronym-dictionary.ts
 * ========================================================================= */

const DICTIONARY_SRC = `./${STATIC_COMPLETION_FILE}`;
const LOAD_TIMEOUT_MS = 20_000;

type PackedLoader = () => Promise<PackedStaticCompletion | null>;

function packedFromWindow(): PackedStaticCompletion | null {
  const value = (globalThis as Record<string, unknown>)[STATIC_COMPLETION_GLOBAL];
  return value && typeof value === 'object' ? (value as PackedStaticCompletion) : null;
}

function injectStaticCompletionScript(): Promise<PackedStaticCompletion | null> {
  const ready = packedFromWindow();
  if (ready !== null) return Promise.resolve(ready);

  if (typeof document === 'undefined' || typeof document.createElement !== 'function') {
    return Promise.resolve(null);
  }

  return new Promise((resolve) => {
    let settled = false;
    const finish = (value: PackedStaticCompletion | null): void => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      resolve(value);
    };

    const timer = setTimeout(() => finish(null), LOAD_TIMEOUT_MS);

    const script = document.createElement('script');
    script.async = false;
    script.src = DICTIONARY_SRC;
    script.addEventListener('load', () => finish(packedFromWindow()));
    script.addEventListener('error', () => finish(null));
    document.head.appendChild(script);
  });
}

let pending: Promise<StaticSource[] | null> | null = null;
let loaded: StaticSource[] | null = null;

/**
 * המקורות, בטעינה עצלה וחד-פעמית — נגררת רק כשבפועל צריך fallback (אין
 * התאמה מהספר הפתוח בקורא), לא בעליית התוסף. כשל מחזיר `null` ואינו נזכר.
 */
export function loadStaticSources(loader: PackedLoader = injectStaticCompletionScript): Promise<StaticSource[] | null> {
  if (loaded) return Promise.resolve(loaded);
  if (pending) return pending;

  pending = (async () => {
    const packed = await loader();
    if (packed === null) return null;
    loaded = buildSources(packed);
    return loaded;
  })()
    .catch(() => null)
    .finally(() => {
      pending = null;
    });

  return pending;
}

/** לבדיקות בלבד. */
export function resetStaticSources(): void {
  loaded = null;
  pending = null;
}

export interface StaticCompletionMatch extends BookCompletionMatch {
  source: string;
}

/**
 * מנסה כל מקור לפי הסדר, מחזירה את ההתאמה הראשונה. אין מיזוג בין מקורות —
 * ברגע שמקור אחד מצא התאמה, השאר לא נבדקים לאותה הקשה.
 */
export function matchStaticCompletion(
  context: TypedContext,
  candidateSources: readonly StaticSource[],
): StaticCompletionMatch | null {
  for (const source of candidateSources) {
    const match = matchAtCursor(source.cache, context, { minStandalonePartial: 2 });
    if (match) return { ...match, source: source.name };
  }
  return null;
}
