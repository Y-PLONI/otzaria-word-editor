/**
 * טעינת מילון ראשי-התיבות כנכס נפרד, בהזרקת `<script>` — אותה תבנית בדיוק
 * כמו `spellcheck-dictionary.ts` (ר' שם ההסבר המלא: IIFE יחיד עם
 * `inlineDynamicImports`, ו-`file://` שחוסם `fetch`). כאן ללא switch נפרד
 * ברצועה — הטעינה קורית בפעם הראשונה שמשתמש מקליד `@` (ר' תכנון), לא
 * בעליית התוסף, ולכן גם משתמש שלא מתייג מקורות אינו טוען את הנכס.
 */
import { ACRONYMS_FILE, ACRONYMS_GLOBAL } from './acronyms-constants';
import { createAcronymDictionary, type AcronymDictionary } from './acronyms';

const DICTIONARY_SRC = `./${ACRONYMS_FILE}`;
const LOAD_TIMEOUT_MS = 20_000;

type PackedLoader = () => Promise<string | null>;

function packedFromWindow(): string | null {
  const value = (globalThis as Record<string, unknown>)[ACRONYMS_GLOBAL];
  return typeof value === 'string' && value.length > 0 ? value : null;
}

function injectAcronymsScript(): Promise<string | null> {
  const ready = packedFromWindow();
  if (ready !== null) return Promise.resolve(ready);

  if (typeof document === 'undefined' || typeof document.createElement !== 'function') {
    return Promise.resolve(null);
  }

  return new Promise((resolve) => {
    let settled = false;
    const finish = (value: string | null): void => {
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

let pending: Promise<AcronymDictionary | null> | null = null;
let loaded: AcronymDictionary | null = null;

/** המילון, בטעינה עצלה וחד-פעמית. כשל אינו נזכר — ניסיון הבא טוען מחדש. */
export function loadAcronymDictionary(
  loader: PackedLoader = injectAcronymsScript,
): Promise<AcronymDictionary | null> {
  if (loaded) return Promise.resolve(loaded);
  if (pending) return pending;

  pending = (async () => {
    const packed = await loader();
    if (packed === null) return null;
    loaded = createAcronymDictionary(packed);
    return loaded;
  })()
    .catch(() => null)
    .finally(() => {
      pending = null;
    });

  return pending;
}

/** לבדיקות בלבד. */
export function resetAcronymDictionary(): void {
  loaded = null;
  pending = null;
}
